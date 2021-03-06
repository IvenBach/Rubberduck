﻿using System;
using System.Collections.Generic;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Parsing.Rewriter
{
    public class RewritingManager : IRewritingManager
    {
        private readonly HashSet<IRewriteSession> _activeCodePaneSessions = new HashSet<IRewriteSession>();
        private readonly HashSet<IRewriteSession> _activeAttributesSessions = new HashSet<IRewriteSession>();
        private readonly IMemberAttributeRecovererWithSettableRewritingManager _memberAttributeRecoverer;

        private readonly IRewriteSessionFactory _sessionFactory;

        private readonly object _invalidationLockObject = new object();

        public RewritingManager(IRewriteSessionFactory sessionFactory, IMemberAttributeRecovererWithSettableRewritingManager memberAttributeRecoverer)
        {
            _sessionFactory = sessionFactory;
            _memberAttributeRecoverer = memberAttributeRecoverer;
            _memberAttributeRecoverer.RewritingManager = this;
        }


        public IRewriteSession CheckOutCodePaneSession()
        {
            var newSession = _sessionFactory.CodePaneSession(TryAllowExclusiveRewrite);
            lock (_invalidationLockObject)
            {
                _activeCodePaneSessions.Add(newSession);
            }

            return newSession;
        }

        public IRewriteSession CheckOutAttributesSession()
        {
            var newSession = _sessionFactory.AttributesSession(TryAllowExclusiveRewrite);
            lock (_invalidationLockObject)
            {
                _activeAttributesSessions.Add(newSession);
            }

            return newSession;
        }

        private bool TryAllowExclusiveRewrite(IRewriteSession rewriteSession)
        {
            lock (_invalidationLockObject)
            {
                if (!IsCurrentlyActive(rewriteSession))
                {
                    return false;
                }

                rewriteSession.Status = RewriteSessionState.RewriteApplied;

                InvalidateAllSessionsInternal();
                if (rewriteSession.TargetCodeKind == CodeKind.CodePaneCode)
                {
                    RequestMemberAttributeRecovery(rewriteSession);
                }
                return true;
            }
        }

        private void RequestMemberAttributeRecovery(IRewriteSession rewriteSession)
        {
            _memberAttributeRecoverer.RecoverCurrentMemberAttributesAfterNextParse(rewriteSession.CheckedOutModules);
        }

        private bool IsCurrentlyActive(IRewriteSession rewriteSession)
        {
            switch (rewriteSession.TargetCodeKind)
            {
                case CodeKind.CodePaneCode:
                    return _activeCodePaneSessions.Contains(rewriteSession);
                case CodeKind.AttributesCode:
                    return _activeAttributesSessions.Contains(rewriteSession);
                default:
                    throw new NotSupportedException(nameof(rewriteSession));
            }
        }

        public void InvalidateAllSessions()
        {
            lock (_invalidationLockObject)
            {
                InvalidateAllSessionsInternal();
            }
        }

        private void InvalidateAllSessionsInternal()
        {
            foreach (var rewriteSession in _activeCodePaneSessions)
            {
                rewriteSession.Status = RewriteSessionState.OtherSessionsRewriteApplied;
            }
            _activeCodePaneSessions.Clear();

            foreach (var rewriteSession in _activeAttributesSessions)
            {
                rewriteSession.Status = RewriteSessionState.OtherSessionsRewriteApplied;
            }
            _activeAttributesSessions.Clear();
        }
    }
}